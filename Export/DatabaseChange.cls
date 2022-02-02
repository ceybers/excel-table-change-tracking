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
'@Folder("Objects")
Option Explicit

Public KeyName As String
Public KeyFK As Long
Public TrackFK As Long
Public CommitFK As Long
Public TableName As String
Public FieldName As String
Public Value As Variant

Public Function ToString() As String
    ToString = KeyFK & "," & TrackFK & "," & TableName & "," & FieldName & "," & Value
End Function

Public Function Pair() As String
    Pair = KeyFK & "," & TableName
End Function
