VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DatabaseChangeFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Loads Schema and KeyName->KeyFK translation table. Provides TranslateChange and TranslateChanges methods."
'@ModuleDescription "Loads Schema and KeyName->KeyFK translation table. Provides TranslateChange and TranslateChanges methods."
'@Folder "DatabaseChanges"
Option Explicit

Private Type TState
    schema As Collection
    Keys As Collection
    LookupTables As LookupTables
End Type

Private this As TState

' Ctor
Private Sub Class_Initialize()
    Set this.schema = New Collection
    Set this.Keys = New Collection
    Set this.LookupTables = New LookupTables
    LoadSchema this.schema
    LoadKeys
    this.LookupTables.Load Access.GetConnection
End Sub

' Public Methods
Public Function Self() As DatabaseChangeFactory
    Set Self = Me
End Function

Public Function TranslateChange(ByVal FieldChange As FieldChange) As DatabaseChange
    Dim SchemaItem As SchemaItem
    
    Dim result As New DatabaseChange
    Set result = New DatabaseChange
    Set SchemaItem = this.schema(FieldChange.ColumnName)
    
    With result
        .KeyName = FieldChange.Key
        .KeyFK = this.Keys(FieldChange.Key)
        .trackFK = -1
        .TableName = SchemaItem.TableName
        .FieldName = SchemaItem.FieldName
        
        If IsObject(FieldChange.After) Then
            Set .Value = FieldChange.After
        Else
            .Value = FieldChange.After
        End If
        
        Dim lkpTable As String
        lkpTable = this.schema(FieldChange.ColumnName).Lookup
        If lkpTable <> vbNullString Then
            .Value = CLng(this.LookupTables.Items(lkpTable).Items(FieldChange.After))
        End If
    End With
    
    Set TranslateChange = result
End Function

Public Sub TranslateChanges(ByVal dbChanges As DatabaseChanges, ByVal FieldChanges As FieldChanges)
    If CheckKeys(FieldChanges) = False Then
        'MsgBox "Orphaned keys detected!"
        Exit Sub
    End If
    
    If CheckColumns(FieldChanges) = False Then
        'MsgBox "Orphaned columns detected!"
        Exit Sub
    End If
    
    Dim change As FieldChange
    For Each change In FieldChanges.Items
        dbChanges.Items.Add TranslateChange(change)
    Next change
    
    dbChanges.UpdateDistinct
End Sub

' Private Methods
Private Sub LoadKeys()
    CollectionClear this.Keys
    
    Dim Recordset As ADODB.Recordset
    Dim sql As String
    Set Recordset = New ADODB.Recordset
    sql = "SELECT * FROM " & KEY_TABLE_NAME
    Recordset.Open Source:=sql, ActiveConnection:=Access.GetConnection

    If Not Recordset.BOF And Not Recordset.EOF Then
        Do While Not Recordset.EOF
            this.Keys.Add Item:=Recordset.fields(0).Value, Key:=Recordset.fields(1).Value
            Recordset.MoveNext
        Loop
    End If
    
    Recordset.Close
    Set Recordset = Nothing
End Sub

Private Function CheckKeys(ByVal FieldChanges As FieldChanges) As Boolean
    CheckKeys = True
    
    Dim fc As FieldChange
    For Each fc In FieldChanges.Items
        If CollectionExists(fc.Key, this.Keys) = False Then
            MsgBox "Orphaned key found = '" & fc.Key & "'", vbExclamation + vbOKOnly
            CheckKeys = False
            Exit Function
        End If
    Next fc
End Function

Private Function CheckColumns(ByVal FieldChanges As FieldChanges) As Boolean
    CheckColumns = True
    
    Dim fc As FieldChange
    For Each fc In FieldChanges.Items
        If CollectionExists2(fc.ColumnName, this.schema) = False Then
            MsgBox "Orphaned column found = '" & fc.ColumnName & "'", vbExclamation + vbOKOnly
            CheckColumns = False
            Exit Function
        End If
    Next fc
End Function

' TODO Refactor
Private Function CollectionExists(ByVal criteria As Variant, ByVal coll As Collection) As Boolean
    Dim v As Variant
    On Error Resume Next
    v = coll(criteria)
    On Error GoTo 0
    CollectionExists = Not IsEmpty(v)
End Function

' TODO Refactor
Private Function CollectionExists2(ByVal criteria As Variant, ByVal coll As Collection) As Boolean
    Dim v As Variant
    On Error Resume Next
    Set v = coll(criteria)
    On Error GoTo 0
    If IsEmpty(v) Then Exit Function
    CollectionExists2 = Not v Is Nothing
End Function
