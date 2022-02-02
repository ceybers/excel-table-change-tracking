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
'@Folder "Objects"
Option Explicit

Public key As String
Public ColumnName As String
Public Before As Variant
Public After As Variant

Public Function Create(ByVal key As String, ByVal ColumnName As String, ByVal Before As Variant, ByVal After As Variant) As FieldChange
    With New FieldChange
        .key = key
        .ColumnName = ColumnName
        .Before = Before
        .After = After
        Set Create = .Self
    End With
End Function

Public Function ToString() As String
    ToString = key & "," & ColumnName & ",'" & Before & "','" & After & "' (" & VarType(After) & ")"
End Function

Public Function Self() As FieldChange
    Set Self = Me
End Function
