VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DatabaseChanges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder("Objects")
Option Explicit

Public Items As Collection
Public DistinctKeys As Scripting.Dictionary
Public DistinctPairs As Scripting.Dictionary

Private Sub Class_Initialize()
    Set Items = New Collection
    Set DistinctKeys = New Scripting.Dictionary
    Set DistinctPairs = New Scripting.Dictionary
End Sub

Public Function Create(ByVal translator As DatabaseChangeFactory, ByVal FieldChanges As FieldChanges) As DatabaseChanges
    Set Create = New DatabaseChanges
    Create.Translate translator, FieldChanges
End Function

Public Sub UpdateDistinct()
    Dim dbChange As DatabaseChange
    For Each dbChange In Items
        DistinctKeys(dbChange.KeyName) = dbChange.KeyName
        DistinctPairs(dbChange.Pair) = dbChange.Pair
    Next dbChange
End Sub

Public Sub Translate(ByVal translator As DatabaseChangeFactory, ByVal FieldChanges As FieldChanges)
    translator.TranslateChanges Me, FieldChanges
End Sub

Public Function Add(ByVal dbChange As DatabaseChange)
    Items.Add dbChange ', Key
End Function
