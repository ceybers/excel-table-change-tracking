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

Private This As TDatabaseChange

' Properties
Public Property Get KeyName() As String
    KeyName = This.KeyName
End Property

Public Property Let KeyName(ByVal RHS As String)
    This.KeyName = RHS
End Property

Public Property Get KeyFK() As Long
    KeyFK = This.KeyFK
End Property

Public Property Let KeyFK(ByVal RHS As Long)
    This.KeyFK = RHS
End Property

Public Property Get trackFK() As Long
    trackFK = This.trackFK
End Property

Public Property Let trackFK(ByVal RHS As Long)
    This.trackFK = RHS
End Property

Public Property Get CommitFK() As Long
    CommitFK = This.CommitFK
End Property

Public Property Let CommitFK(ByVal RHS As Long)
    This.CommitFK = RHS
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

Public Property Get Value() As Variant
    If IsObject(This.Value) Then
        Set Value = This.Value
    Else
        Value = This.Value
    End If
End Property

Public Property Let Value(ByVal RHS As Variant)
    This.Value = RHS
End Property

Public Property Set Value(ByVal RHS As Variant)
    Set This.Value = RHS
End Property

' Public Methods
Public Function ToString() As String
    ToString = KeyFK & "," & trackFK & "," & TableName & "," & FieldName & "," & Value
End Function

Public Function Pair() As String
    Pair = KeyFK & "," & TableName
End Function
