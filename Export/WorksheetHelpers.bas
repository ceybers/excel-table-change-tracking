Attribute VB_Name = "WorksheetHelpers"
'@Folder "HelperFunctions"
Option Explicit
Option Private Module

Public Function AddOrGetWorksheet(ByVal worksheetName As String) As Worksheet
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim prevWS As Worksheet
    
    Set wb = ActiveWorkbook
    Set prevWS = ActiveSheet
    
    On Error Resume Next
    Set ws = wb.Worksheets(worksheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = worksheetName
    End If
    
    prevWS.Activate
    
    ws.Visible = xlSheetVeryHidden
    
    Set AddOrGetWorksheet = ws
End Function
