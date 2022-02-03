VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FieldChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder "FieldChanges"
Option Explicit

Private Type TFieldChange
    Key As String
    ColumnName As String
    Before As Variant
    After As Variant
End Type

Private this As TFieldChange

' Properties
Public Property Get Key() As String
    Key = this.Key
End Property

Public Property Let Key(ByVal RHS As String)
    this.Key = RHS
End Property

Public Property Get ColumnName() As String
    ColumnName = this.ColumnName
End Property

Public Property Let ColumnName(ByVal RHS As String)
    this.ColumnName = RHS
End Property

Public Property Get Before() As Variant
    If IsObject(this.Before) Then
        Set Before = this.Before
    Else
        Before = this.Before
    End If
End Property

Public Property Let Before(ByVal RHS As Variant)
    this.Before = RHS
End Property

Public Property Set Before(ByVal RHS As Variant)
    Set this.Before = RHS
End Property

Public Property Get After() As Variant
    If IsObject(this.After) Then
        Set After = this.After
    Else
        After = this.After
    End If
End Property

Public Property Let After(ByVal RHS As Variant)
    this.After = RHS
End Property

Public Property Set After(ByVal RHS As Variant)
    Set this.After = RHS
End Property

' Public Methods
Public Function Create(ByVal Key As String, ByVal ColumnName As String, ByVal Before As Variant, ByVal After As Variant) As FieldChange
    With New FieldChange
        .Key = Key
        .ColumnName = ColumnName
        .Before = Before
        .After = After
        Set Create = .Self
    End With
End Function

Public Function ToString() As String
    ToString = Key & "," & ColumnName & ",'" & Before & "','" & After & "' (" & VarType(After) & ")"
End Function

Public Function Self() As FieldChange
    Set Self = Me
End Function
