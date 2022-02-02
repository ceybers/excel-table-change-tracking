VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SchemaItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Objects")
Option Explicit

Private Type TSchemaItem
    ColumnName As String
    TableName As String
    FieldName As String
    VarType As Long
End Type

Private this As TSchemaItem

Public Property Get ColumnName() As String
    ColumnName = this.ColumnName
End Property

Public Property Get TableName() As String
    TableName = this.TableName
End Property

Public Property Let TableName(ByVal RHS As String)
    this.TableName = RHS
End Property

Public Property Get FieldName() As String
    FieldName = this.FieldName
End Property

Public Property Let FieldName(ByVal RHS As String)
    this.FieldName = RHS
End Property

Public Property Get VarType() As Long
    VarType = this.VarType
End Property

Public Property Let VarType(ByVal RHS As Long)
    this.VarType = RHS
End Property

Public Sub SetFromString(ByVal Payload As String)
    Dim SplitPayload As Variant
    SplitPayload = Split(Payload, ",")
    With this
        .ColumnName = SplitPayload(0)
        .TableName = SplitPayload(1)
        .FieldName = SplitPayload(2)
        .VarType = SplitPayload(3)
    End With
End Sub

Public Sub SetValues(ColumnName As String, TableName As String, FieldName As String, VarType As Long)
    With this
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
    ToString = this.ColumnName & "," & this.TableName & "," & this.FieldName & "," & this.VarType
End Function