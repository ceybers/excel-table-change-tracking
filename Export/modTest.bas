Attribute VB_Name = "modTest"
'@Folder("VBAProject")
Option Explicit

Public Sub AAATest()
    Dim lkp As LookupTable
    Set lkp = New LookupTable
    lkp.TableName = "lkpYesNoConfirmNA"
    lkp.Load Access.GetConnection
    
    Debug.Print lkp.Items.Count
    
    Access.GetConnection DoClose:=True
End Sub

Public Sub Test()
    Dim conn As ADODB.Connection
    
    Set conn = Access.GetConnection
    
    Dim rs As ADODB.Recordset
    'Set rs = New ADODB.Recordset
    'Call rs.Open("SELECT ID, Title FROM metaCommits", conn, adOpenStatic, adLockReadOnly)
    
    'Debug.Print rs.RecordCount
    'Dim v As Variant
    'v = rs.GetRows
    
    Set rs = conn.OpenSchema(adSchemaTables)
    'Debug.Print rs.RecordCount
    Dim v As Variant
    v = rs.GetRows
    
    Dim i As Long
    For i = 0 To UBound(v)
        Debug.Print v(2, i)
    Next i
    
    Access.GetConnection DoClose:=True
End Sub
