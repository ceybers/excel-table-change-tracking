VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LookupTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Access")
Option Explicit

Private Type TState
     TableName As String
     Item As Scripting.Dictionary
End Type

Private this As TState

Public Property Get TableName() As String
    TableName = this.TableName
End Property

Public Property Let TableName(ByVal RHS As String)
    this.TableName = RHS
End Property

'@DefaultMember
Public Property Get Items() As Scripting.Dictionary
Attribute Items.VB_UserMemId = 0
    Set Items = this.Item
End Property

Public Property Set Items(ByVal RHS As Scripting.Dictionary)
    Set this.Item = RHS
End Property

Private Sub Class_Initialize()
    Set Items = New Scripting.Dictionary
End Sub

Public Sub Load(ByVal db As ADODB.Connection)
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open this.TableName, db, adOpenStatic, adLockReadOnly
    
    Dim rows As Variant
    rows = rs.GetRows
    
    this.Item.RemoveAll
    Dim i As Long
    For i = 0 To UBound(rows, 2)
        this.Item.Add Key:=rows(1, i), Item:=rows(0, i)
    Next i
    
    rs.Close
    Set rs = Nothing
End Sub
