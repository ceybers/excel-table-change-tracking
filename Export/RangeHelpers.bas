Attribute VB_Name = "RangeHelpers"
'@Folder("HelperFunctions")
Option Explicit

Public Sub AppendRange(ByVal rangeToAppend As Range, ByRef unionRange As Range)
    If rangeToAppend Is Nothing Then Exit Sub
    
    If unionRange Is Nothing Then
        Set unionRange = rangeToAppend
        Exit Sub
    End If
    
    If Not rangeToAppend.Parent Is unionRange.Parent Then Exit Sub
    
    Set unionRange = Application.Union(unionRange, rangeToAppend)
End Sub
