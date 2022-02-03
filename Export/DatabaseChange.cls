VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DatabaseChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder "DatabaseChanges"
Option Explicit

Private Type TDatabaseChange
    KeyName As String
    KeyFK As Long
    trackFK As Long
    CommitFK As Long
    TableName As String
    FieldName As String
    Value As Variant
End Type

Private this As TDatabaseChange

' Properties
Public Property Get KeyName() As String
    KeyName = this.KeyName
End Property

Public Property Let KeyName(ByVal RHS As String)
    this.KeyName = RHS
End Property

Public Property Get KeyFK() As Long
    KeyFK = this.KeyFK
End Property

Public Property Let KeyFK(ByVal RHS As Long)
    this.KeyFK = RHS
End Property

Public Property Get trackFK() As Long
    trackFK = this.trackFK
End Property

Public Property Let trackFK(ByVal RHS As Long)
    this.trackFK = RHS
End Property

Public Property Get CommitFK() As Long
    CommitFK = this.CommitFK
End Property

Public Property Let CommitFK(ByVal RHS As Long)
    this.CommitFK = RHS
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

Public Property Get Value() As Variant
    If IsObject(this.Value) Then
        Set Value = this.Value
    Else
        Value = this.Value
    End If
End Property

Public Property Let Value(ByVal RHS As Variant)
    this.Value = RHS
End Property

Public Property Set Value(ByVal RHS As Variant)
    Set this.Value = RHS
End Property

' Public Methods
Public Function ToString() As String
    ToString = KeyFK & "," & trackFK & "," & TableName & "," & FieldName & "," & Value
End Function

Public Function Pair() As String
    Pair = KeyFK & "," & TableName
End Function
