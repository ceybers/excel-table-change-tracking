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

