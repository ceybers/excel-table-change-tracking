VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SchemaItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("DatabaseChanges")
Option Explicit

Private Type TSchemaItem
    ColumnName As String
    TableName As String
    FieldName As String
    VarType As Long
End Type

Private This As TSchemaItem

' Properties
Public Property Get ColumnName() As String
    ColumnName = This.ColumnName
End Property

Public Property Get TableName() As String
    TableName = This.TableName
End Property

Public Property Let TableName(ByVal RHS As String)
    This.TableName = RHS
End Property

Public Property Get FieldName() As String
    FieldName = This.FieldName
End Property

Public Property Let FieldName(ByVal RHS As String)
    This.FieldName = RHS
End Property

Public Property Get VarType() As Long
    VarType = This.VarType
End Property

Public Property Let VarType(ByVal RHS As Long)
    This.VarType = RHS
End Property

' Public Methods
Public Sub SetFromString(ByVal Payload As String)
    Dim SplitPayload As Variant
    SplitPayload = Split(Payload, ",")
    With This
        .ColumnName = SplitPayload(0)
        .TableName = SplitPayload(1)
        .FieldName = SplitPayload(2)
        .VarType = SplitPayload(3)
    End With
End Sub

Public Sub SetValues(ByVal ColumnName As String, ByVal TableName As String, ByVal FieldName As String, ByVal VarType As Long)
    With This
        .ColumnName = ColumnName
        .TableName = TableName
        .FieldName = FieldName
        .VarType = VarType
    End With
End Sub

Public Function Self() As SchemaItem
    Set Self = Me
End Function

Public Function ToString() As String
    ToString = This.ColumnName & "," & This.TableName & "," & This.FieldName & "," & This.VarType
End Function
