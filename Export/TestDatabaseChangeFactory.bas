Attribute VB_Name = "TestDatabaseChangeFactory"
'@Folder("Objects")
Option Explicit
Option Private Module

Public Sub Test()
    Dim dbchgfact As DatabaseChangeFactory
    Set dbchgfact = New DatabaseChangeFactory
    'dbchgfact.Load
    
    Dim Changes As New FieldChanges
    Set Changes = New FieldChanges
    Changes.GenerateTestFieldChanges
    
    Dim dbChanges As New Collection
    'Set dbChanges = dbchgfact.TranslateChanges(Changes)
    
    Dim dbChange As DatabaseChange
    For Each dbChange In dbChanges
        Debug.Print dbChange.ToString
    Next dbChange
End Sub

