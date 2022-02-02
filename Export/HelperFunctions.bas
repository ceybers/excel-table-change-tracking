Attribute VB_Name = "HelperFunctions"
'@Folder "HelperFunctions"
Option Explicit
Option Private Module

Public Function AddOrGetWorksheet(ByVal worksheetName As String) As Worksheet
    Dim wb As Workbook
    Dim ws As Worksheet
    
    Set wb = ActiveWorkbook
    On Error Resume Next
    Set ws = wb.Worksheets(worksheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = worksheetName
    End If
    
    Set AddOrGetWorksheet = ws
End Function
