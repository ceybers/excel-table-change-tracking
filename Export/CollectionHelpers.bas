Attribute VB_Name = "CollectionHelpers"
'@Folder "HelperFunctions"
Option Explicit
Option Private Module

Public Function CollectionClear(ByVal coll As Collection) As Boolean
    If coll Is Nothing Then Exit Function
    If coll.Count = 0 Then Exit Function
    
    Do While coll.Count > 0
        coll.Remove 1
    Loop
    
    CollectionClear = True
End Function

Public Function CollectionExists(ByVal coll As Collection, ByVal criteria As Variant) As Boolean
    Dim v As Variant
    For Each v In coll
        If v = criteria Then
            CollectionExists = True
            Exit Function
        End If
    CollectionExists = False
End Function

