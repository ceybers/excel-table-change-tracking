VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LookupTables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "LookupTables"
Option Explicit

Private Type TLookupTables
    Items As Scripting.Dictionary
End Type

Private this As TLookupTables

Public Property Get Items() As Scripting.Dictionary
    Set Items = this.Items
End Property

Private Sub Class_Initialize()
    Set this.Items = New Scripting.Dictionary
End Sub

Public Sub Load(ByVal db As ADODB.Connection)
    Dim rs As ADODB.Recordset
    Set rs = db.OpenSchema(adSchemaTables)

    Dim rows As Variant
    rows = rs.GetRows
    
    Dim lkpTable As LookupTable
    Dim tblName As String
    Dim i As Long
    For i = 0 To UBound(rows, 2)
        tblName = rows(2, i)
        If tblName Like "lkp*" Then
            Set lkpTable = New LookupTable
            lkpTable.TableName = tblName
            lkpTable.Load db
            Items.Add Key:=tblName, Item:=lkpTable
        End If
    Next i
End Sub
