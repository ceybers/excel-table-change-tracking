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
'@Folder "DatabaseChanges"
Option Explicit

Private Type TDatabaseChanges
    Items As Collection
    DistinctKeys As Scripting.Dictionary
    DistinctPairs As Scripting.Dictionary
End Type

Private This As TDatabaseChanges

' Properties
Public Property Get Items() As Collection
    Set Items = This.Items
End Property

Public Property Get DistinctKeys() As Scripting.Dictionary
    Set DistinctKeys = This.DistinctKeys
End Property

Public Property Get DistinctPairs() As Scripting.Dictionary
    Set DistinctPairs = This.DistinctPairs
End Property

' Ctor
Private Sub Class_Initialize()
    Set This.Items = New Collection
    Set This.DistinctKeys = New Scripting.Dictionary
    Set This.DistinctPairs = New Scripting.Dictionary
End Sub

' Public Methods
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

Public Sub Add(ByVal dbChange As DatabaseChange)
    Items.Add dbChange 'TODO Use a key?
End Sub
