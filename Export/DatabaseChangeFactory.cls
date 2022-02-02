VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DatabaseChangeFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "hi"
'@ModuleDescription "Loads Schema and KeyName->KeyFK translation table. Provides TranslateChange and TranslateChanges methods."
'@Folder("Objects")
Option Explicit

Private Type TState
    Schema As Collection
    Keys As Collection
End Type

Private this As TState

Public Function Self() As DatabaseChangeFactory
    Set Self = Me
End Function

Private Sub Class_Initialize()
    Set this.Schema = New Collection
    Set this.Keys = New Collection
    LoadSchema
    LoadKeys
End Sub

Private Sub LoadSchema()
    Dim InputLine As String
    
    Set this.Schema = New Collection
    
    Open SCHEMA_PATH For Input As #1
    Do Until EOF(1)
        Line Input #1, InputLine
        With New SchemaItem
            .SetFromString InputLine
            this.Schema.Add item:=.Self, key:=.ColumnName
        End With
    Loop
    Close #1
End Sub

Private Sub LoadKeys()
    Dim Connection As ADODB.Connection
    Dim Recordset As ADODB.Recordset
    Dim sql As String
    
    CollectionClear this.Keys
    
    Set Connection = New ADODB.Connection
    Connection.Open ConnectionString:=Access.GetConnectionString
    
    Set Recordset = New ADODB.Recordset
    sql = "SELECT * FROM " & KEY_TABLE_NAME
    Recordset.Open Source:=sql, ActiveConnection:=Connection

    If Not Recordset.BOF And Not Recordset.EOF Then
        Do While Not Recordset.EOF
            this.Keys.Add item:=Recordset.fields(0).Value, key:=Recordset.fields(1).Value
            Recordset.MoveNext
        Loop
    End If
    
    Recordset.Close
    Set Recordset = Nothing
    Set Connection = Nothing
End Sub

Public Function TranslateChange(ByVal FieldChange As FieldChange) As DatabaseChange
    Dim SchemaItem As SchemaItem
    
    Set TranslateChange = New DatabaseChange
    Set SchemaItem = this.Schema(FieldChange.ColumnName)
    
    With TranslateChange
        .KeyName = FieldChange.key
        .KeyFK = this.Keys(FieldChange.key)
        .TrackFK = -1
        .TableName = SchemaItem.TableName
        .FieldName = SchemaItem.FieldName
        
        If IsObject(FieldChange.After) Then
            Set .Value = FieldChange.After
        Else
            .Value = FieldChange.After
        End If
    End With
End Function

Public Sub TranslateChanges(ByVal dbChanges As DatabaseChanges, ByVal FieldChanges As FieldChanges)
    Dim change As FieldChange
    For Each change In FieldChanges.Items
        dbChanges.Items.Add TranslateChange(change)
    Next change
    
    dbChanges.UpdateDistinct
End Sub
