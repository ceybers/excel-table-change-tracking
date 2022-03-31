Attribute VB_Name = "Schema"
'@Folder("DatabaseChanges")
Option Explicit

Public Sub LoadSchema(ByVal coll As Collection)
    Debug.Assert Not coll Is Nothing
    
    CollectionHelpers.CollectionClear coll
    
    Dim InputLine As String
    
    Open SCHEMA_PATH For Input As #1
    Do Until EOF(1)
        Line Input #1, InputLine
        With New SchemaItem
            .SetFromString InputLine
            coll.Add Item:=.Self, Key:=.ColumnName
        End With
    Loop
    Close #1
End Sub
