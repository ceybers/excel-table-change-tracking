VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FieldChanges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "FieldChanges"
Option Explicit

Private Type TFieldChanges
    Working As ListObject
    Before As Worksheet
    Items As Collection
    BeforeArray As Variant
End Type

Private this As TFieldChanges

' Properties
'@DefaultMember
Public Property Get Items() As Collection
Attribute Items.VB_UserMemId = 0
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

' Ctor
Private Sub Class_Initialize()
    Set this.Items = New Collection
End Sub

' Public Methods
Public Function IsValid() As Boolean
    If Working Is Nothing Then Exit Function
    If Before Is Nothing Then Exit Function
    IsValid = True
End Function

Public Function Self() As FieldChanges
    Set Self = Me
End Function

Public Sub Compare()
    CollectionClear this.Items
    
    Dim lc As ListColumn
    For Each lc In this.Working.ListColumns
        If lc.DataBodyRange.HasFormula = False And lc.Index <> 1 Then ' TODO Allow key columns in other column index
            CompareFieldsInColumn lc
        End If
    Next lc
End Sub

' Private Methods
Private Sub CompareFieldsInColumn(ByVal lc As ListColumn) 'as Variant
    Dim AfterArray As Variant
    
    AfterArray = lc.DataBodyRange.Value2
    
    Dim ThisVarType As Long
    ThisVarType = vbString
    If InStr(lc.DataBodyRange.Cells(1, 1).NumberFormat, "y") > 0 Then
        ThisVarType = vbDate
    End If
    
    Dim i As Long
    For i = 1 To UBound(AfterArray, 1)
        CompareField AfterArray, i, lc.Index, ThisVarType
    Next i
End Sub

Private Sub CompareField(ByRef AfterArray As Variant, ByVal Row As Long, ByVal Column As Long, ByVal VarType As Long)
    Dim newColor As Long
    Dim Before As Variant
    Dim After As Variant

    'If Not IsValid Then MsgBox "FieldChanges not valid"
    Before = this.BeforeArray(Row + 1, Column)
    After = AfterArray(Row, 1)
    
    If IsError(Before) Or IsError(After) Then
        newColor = RGB(255, 255, 64)
    ElseIf Before <> After Then
        With New FieldChange
            .Key = this.BeforeArray(Row + 1, 1)
            .ColumnName = this.BeforeArray(1, Column)
            .Before = Before
            If VarType = vbDate And Not IsEmpty(After) Then
                .After = CDate(After)
            Else
                .After = After
            End If
            this.Items.Add .Self
        End With
        newColor = RGB(64, 255, 128)
    Else
        newColor = xlNone
    End If
    
    'this.Working.DataBodyRange.Cells(Row, Column).Interior.Color = newColor
End Sub

Public Sub DEBUGPrintChanges()
    Dim chg As FieldChange
    For Each chg In this.Items
        Debug.Print chg.ToString
    Next chg
End Sub
